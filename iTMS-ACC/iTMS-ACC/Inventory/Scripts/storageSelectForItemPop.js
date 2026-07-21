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
		return frm && frm.elements ? frm.elements[name] : null;
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

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, String(value == null ? "" : value));
		}
	}

	function table() {
		return document.getElementById("tblData");
	}

	function addExistingStoreRow(index, storeName) {
		var row = table().insertRow(-1);
		var cell = row.insertCell(-1);
		cell.textContent = String(index);
		cell.className = "ExcelSerial";
		cell.align = "center";
		cell = row.insertCell(-1);
		cell.textContent = storeName;
		cell.className = "ExcelDisplayCell";
	}

	function resolveParentXml() {
		ensureCompat();
		parentDoc = xmlDocument(window.dialogArguments);
		parentRoot = xmlRoot(window.dialogArguments);
		if (!parentDoc && window.opener && window.opener.StoreData) {
			parentDoc = xmlDocument(window.opener.StoreData);
			parentRoot = xmlRoot(window.opener.StoreData);
		}
		if (!parentDoc) {
			parentDoc = xmlDocument(document.getElementById("ItemData"));
			parentRoot = xmlRoot(document.getElementById("ItemData"));
		}
	}

	window.fnInit = function () {
		resolveParentXml();
		elementChildren(parentRoot, "STOREDET").forEach(function (store, index) {
			addExistingStoreRow(index + 1, getAttr(store, "STORE"));
		});
	};

	window.CheckSubmit = function () {
		var select = field("selStore");
		var node;
		if (!select || select.selectedIndex < 0 || select.options[select.selectedIndex].value === "S") {
			alert("Select Store");
			if (select) {
				select.focus();
			}
			return false;
		}
		resolveParentXml();
		if (!parentDoc || !parentRoot) {
			return false;
		}
		node = parentDoc.createElement("STOREDET");
		setAttr(node, "UNITSTORE", select.options[select.selectedIndex].value);
		setAttr(node, "STORE", select.options[select.selectedIndex].text);
		parentRoot.appendChild(node);
		window.returnValue = parentRoot;
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(parentRoot);
		}
		window.close();
		return false;
	};
}(window, document));
