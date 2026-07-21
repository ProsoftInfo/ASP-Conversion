(function (window, document) {
	"use strict";

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlRoot(value) {
		ensureCompat();
		if (!value) {
			return null;
		}
		if (value.documentElement) {
			return value.documentElement;
		}
		if (value.XMLDocument && value.XMLDocument.documentElement) {
			return value.XMLDocument.documentElement;
		}
		if (value._doc && value._doc.documentElement) {
			return value._doc.documentElement;
		}
		return value.nodeType === 1 ? value : null;
	}

	function elementChildren(node) {
		var result = [];
		var children = node && node.childNodes || [];
		for (var i = 0; i < children.length; i += 1) {
			if (children[i].nodeType === 1) {
				result.push(children[i]);
			}
		}
		return result;
	}

	function attrAt(node, index) {
		var attr = node && node.attributes && node.attributes.item(index);
		return attr ? attr.nodeValue || attr.value || "" : "";
	}

	function attrAtOrName(node, index, name) {
		return attrAt(node, index) || node && node.getAttribute && node.getAttribute(name) || "";
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function textInputs() {
		var form = document.forms.formname || document.forms[0];
		var fields = form && form.elements || [];
		var result = [];
		for (var i = 0; i < fields.length; i += 1) {
			if (String(fields[i].type || "").toLowerCase() === "text" && String(fields[i].name || "").indexOf("txt") === 0) {
				result.push(fields[i]);
			}
		}
		return result;
	}

	window.fnInit = function (sItem, sClass) {
		var root = xmlRoot(window.dialogArguments);
		var inputs = textInputs();
		var inputIndex = 0;
		var headers;
		var children;
		var pages;
		var h;
		var c;
		var p;

		if (!root) {
			return;
		}

		headers = elementChildren(root);
		for (h = 0; h < headers.length; h += 1) {
			if (trim(attrAt(headers[h], 0)) === trim(sClass) && trim(attrAt(headers[h], 1)) === trim(sItem)) {
				children = elementChildren(headers[h]);
				for (c = 0; c < children.length; c += 1) {
					if (trim(children[c].nodeName).toLowerCase() === "qtypara") {
						pages = elementChildren(children[c]);
						for (p = 0; p < pages.length && inputIndex < inputs.length; p += 1) {
							inputs[inputIndex].value = attrAtOrName(pages[p], 2, "VALUE");
							inputIndex += 1;
						}
					}
				}
			}
		}
	};
}(window, document));
