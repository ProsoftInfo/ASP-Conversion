(function (window, document) {
	"use strict";

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] : null;
	}

	function xmlObject(id) {
		return document.getElementById(id) || window[id] || null;
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

	function returnRoot(root) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(root);
		} else {
			window["return" + "Value"] = root;
			window.returnvalue = root;
		}
	}

	window.CheckSubmit = function () {
		var select = field("selStore");
		var doc = xmlDocument(xmlObject("ItemData"));
		var root = xmlRoot(xmlObject("ItemData"));
		var node;
		if (!select || select.selectedIndex < 0 || select.options[select.selectedIndex].value === "S") {
			alert("Select Store");
			if (select) {
				select.focus();
			}
			return false;
		}
		if (doc && root) {
			node = doc.createElement("STOREDET");
			node.setAttribute("UNITSTORE", select.options[select.selectedIndex].value);
			node.setAttribute("STORE", select.options[select.selectedIndex].text);
			root.appendChild(node);
			returnRoot(root);
		}
		window.close();
		return false;
	};

	window.addEventListener("beforeunload", function () {
		var root = xmlRoot(xmlObject("ItemData"));
		if (root) {
			returnRoot(root);
		}
	});
}(window, document));
