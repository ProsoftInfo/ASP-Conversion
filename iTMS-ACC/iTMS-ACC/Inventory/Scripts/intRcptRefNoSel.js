(function (window, document) {
	"use strict";

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

	function xmlDocument(value) {
		ensureCompat();
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

	window.CheckSubmit = function () {
		var select = field("selReference");
		var data = document.getElementById("TempData") || window.TempData;
		var doc = xmlDocument(data);
		var root = xmlRoot(data);
		var child;
		if (!select || select.selectedIndex === 0) {
			alert("Select Reference No");
			if (select) {
				select.focus();
			}
			return false;
		}
		while (root && root.firstChild) {
			root.removeChild(root.firstChild);
		}
		child = doc.createElement("Child");
		child.setAttribute("value", select.options[select.selectedIndex].value);
		child.setAttribute("name", select.options[select.selectedIndex].text);
		root.appendChild(child);
		window.returnValue = root;
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(root);
		}
		window.close();
		return false;
	};
}(window, document));
