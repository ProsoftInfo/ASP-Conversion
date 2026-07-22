(function (window, document) {
	"use strict";

	var parentDoc = null;
	var parentRoot = null;
	var itemCode = "";
	var classCode = "";

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

	function modalArgs() {
		ensureCompat();
		return window["dialog" + "Arguments"] || null;
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

	function resolveParentXml() {
		var args = modalArgs();
		parentDoc = xmlDocument(args);
		parentRoot = xmlRoot(args);
	}

	function matchingItem() {
		var found = null;
		elementChildren(parentRoot).some(function (node) {
			if (getAttr(node, "ICODE") === itemCode && getAttr(node, "CCODE") === classCode) {
				found = node;
				return true;
			}
			return false;
		});
		return found;
	}

	function stageQuantity(index) {
		var qtyNode = document.getElementById("idStageQty" + index);
		return qtyNode ? trim(qtyNode.textContent || qtyNode.innerText) : "";
	}

	function returnValueAndClose() {
		if (parentRoot) {
			if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
				window.ITMSModernCompat.returnModalValue(parentRoot);
			} else {
				window["return" + "Value"] = parentRoot;
			}
		}
		window.close();
	}

	window.fnInit = function (classValue, itemValue) {
		var count = Number(field("hiCtr") && field("hiCtr").value || 0);
		var item;
		var index = 0;
		classCode = trim(classValue);
		itemCode = trim(itemValue);
		if (!count) {
			return false;
		}
		resolveParentXml();
		item = matchingItem();
		elementChildren(item, "STAGE").forEach(function (stage) {
			index += 1;
			if (index <= count && field("txtA" + index)) {
				field("txtA" + index).value = getAttr(stage, "IVALUE");
			}
		});
		return false;
	};

	window.CheckSubmit = function () {
		var count = Number(field("hiCtr") && field("hiCtr").value || 0);
		var item;
		var index = 0;
		var valueField;
		if (!count) {
			return false;
		}
		for (var i = 1; i <= count; i += 1) {
			valueField = field("txtA" + i);
			if (!valueField || trim(valueField.value) === "") {
				alert("Enter Value");
				if (valueField) {
					valueField.select();
				}
				return false;
			}
		}
		resolveParentXml();
		item = matchingItem();
		elementChildren(item, "STAGE").forEach(function (stage) {
			index += 1;
			valueField = field("txtA" + index);
			if (index <= count && valueField) {
				setAttr(stage, "IVALUE", valueField.value);
				setAttr(stage, "IQTY", stageQuantity(index));
			}
		});
		returnValueAndClose();
		return false;
	};
}(window, document));
