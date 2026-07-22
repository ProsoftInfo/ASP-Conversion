(function (window, document) {
	"use strict";

	var parentDoc = null;
	var parentRoot = null;

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] || document.getElementById(name) : document.getElementById(name);
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function xmlDocument(value) {
		if (!value) {
			return null;
		}
		if (value.XMLDocument) {
			return value.XMLDocument;
		}
		if (value._doc) {
			return value._doc;
		}
		if (value.nodeType === 9) {
			return value;
		}
		return value.ownerDocument || null;
	}

	function xmlRoot(value) {
		if (!value) {
			return null;
		}
		return value.documentElement || value.XMLDocument && value.XMLDocument.documentElement || value._doc && value._doc.documentElement || value;
	}

	function modalArgs() {
		ensureCompat();
		return window["dialog" + "Arguments"] || null;
	}

	function parseXml(text) {
		return new DOMParser().parseFromString(text || "<Root/>", "text/xml");
	}

	function elementChildren(node, name) {
		var result = [];
		var wanted = name && String(name).toLowerCase();
		var children = node && node.childNodes || [];
		for (var i = 0; i < children.length; i += 1) {
			if (children[i].nodeType === 1 && (!wanted || String(children[i].nodeName).toLowerCase() === wanted)) {
				result.push(children[i]);
			}
		}
		return result;
	}

	function getAttr(node, name) {
		return trim(node && node.getAttribute ? node.getAttribute(name) : "");
	}

	function selectedOption(name) {
		var select = field(name);
		if (!select || !select.options || select.selectedIndex < 0) {
			return { value: "", text: "" };
		}
		return {
			value: select.options[select.selectedIndex].value,
			text: select.options[select.selectedIndex].text
		};
	}

	function clearOptions(select) {
		while (select && select.options && select.options.length) {
			select.remove(0);
		}
	}

	function addOption(select, text, value) {
		var option = document.createElement("option");
		option.text = text;
		option.value = value;
		select.add(option);
	}

	function resolveParent() {
		ensureCompat();
		var args = modalArgs();
		parentDoc = xmlDocument(args) || parentDoc;
		parentRoot = xmlRoot(args) || parentRoot;
		if (parentRoot) {
			parentRoot.setAttribute("Confirm", "N");
		}
	}

	function returnRoot() {
		if (!parentRoot) {
			return;
		}
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(parentRoot);
		} else {
			window["return" + "Value"] = parentRoot;
			window.returnvalue = parentRoot;
		}
	}

	window.init = function () {
		resolveParent();
		return false;
	};

	window.FinalSubmit = function () {
		var invType = selectedOption("cmbInvType");
		var saleType = selectedOption("cmbSaletype");
		var pos = selectedOption("cmbPOS");
		var node;
		resolveParent();
		if (trim(invType.value) === "0" || !trim(invType.value)) {
			alert("Select Invoice Type");
			if (field("cmbInvType")) {
				field("cmbInvType").focus();
			}
			return false;
		}
		if (trim(saleType.value) === "0" || !trim(saleType.value)) {
			alert("Select Invoice Type");
			if (field("cmbSaletype")) {
				field("cmbSaletype").focus();
			}
			return false;
		}
		if (trim(field("hPOSMandatory") && field("hPOSMandatory").value) === "Y" && Number(field("hCountPOS") && field("hCountPOS").value || 0) === 0) {
			alert("Create POS");
			return false;
		}
		node = parentDoc.createElement("SALINV");
		node.setAttribute("InvType", invType.value);
		node.setAttribute("SalType", saleType.value);
		node.setAttribute("InvTypeName", invType.text);
		node.setAttribute("SalTypeName", saleType.text);
		if (trim(field("hPOSMandatory") && field("hPOSMandatory").value) === "Y") {
			node.setAttribute("POS", pos.value);
			node.setAttribute("POSName", pos.text);
		} else {
			node.setAttribute("POS", "");
			node.setAttribute("POSName", "");
		}
		parentRoot.appendChild(node);
		parentRoot.setAttribute("Confirm", "Y");
		parentRoot.setAttribute("POConfirm", "N");
		returnRoot();
		window.close();
		return false;
	};

	window.PopTaxType = function () {
		var invType = selectedOption("cmbInvType").value;
		var saleType = field("cmbSaletype");
		var xhr = new XMLHttpRequest();
		xhr.open("GET", "../../Sales/Transaction/GetSaleType.asp?InvType=" + encodeURIComponent(invType), true);
		xhr.onreadystatechange = function () {
			var xml;
			var root;
			if (xhr.readyState !== 4) {
				return;
			}
			xml = xhr.responseXML && xhr.responseXML.documentElement ? xhr.responseXML : parseXml(xhr.responseText);
			root = xmlRoot(xml);
			if (!root || !elementChildren(root).length) {
				if (trim(xhr.responseText)) {
					alert(xhr.responseText);
				}
				return;
			}
			clearOptions(saleType);
			addOption(saleType, "Select", "0");
			elementChildren(root).forEach(function (node) {
				addOption(saleType, getAttr(node, "Name"), getAttr(node, "Code"));
			});
			saleType.selectedIndex = 0;
		};
		xhr.send(null);
		return false;
	};

	window.addEventListener("beforeunload", returnRoot);
}(window, document));
