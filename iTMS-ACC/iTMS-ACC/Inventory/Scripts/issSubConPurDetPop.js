(function (window, document) {
	"use strict";

	var parentDoc = null;
	var parentRoot = null;
	var purTypeDoc = null;
	var purTypeRoot = null;

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

	function selectedValue(name) {
		var select = field(name);
		return select && select.options && select.selectedIndex >= 0 ? select.options[select.selectedIndex].value : select && select.value || "";
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

	function resolvePurTypeData(xml) {
		purTypeDoc = xml || xmlDocument(xmlObject("PurTypeData")) || document.implementation.createDocument("", "Root", null);
		purTypeRoot = xmlRoot(purTypeDoc) || purTypeDoc.documentElement;
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

	function loadPurchaseTypes(callback) {
		var unit = field("selUnit") ? field("selUnit").value : "";
		var xhr = new XMLHttpRequest();
		xhr.open("GET", "../../purchase/transaction/XMLGetPurchaseType.asp?ForUnit=" + encodeURIComponent(unit), true);
		xhr.onreadystatechange = function () {
			if (xhr.readyState !== 4) {
				return;
			}
			resolvePurTypeData(xhr.responseXML && xhr.responseXML.documentElement ? xhr.responseXML : parseXml(xhr.responseText));
			if (callback) {
				callback();
			}
		};
		xhr.send(null);
	}

	function appendMatchingPurAcc() {
		var wanted = selectedValue("cmbPurType");
		var found = false;
		if (!wanted) {
			alert("Select Purchase Type");
			if (field("cmbPurType")) {
				field("cmbPurType").focus();
			}
			return;
		}
		elementChildren(purTypeRoot).forEach(function (unitNode) {
			elementChildren(unitNode, "PURACC").forEach(function (purAcc) {
				if (!found && getAttr(purAcc, "PurType") === wanted) {
					parentRoot.appendChild(parentDoc.importNode ? parentDoc.importNode(purAcc, true) : purAcc.cloneNode(true));
					found = true;
				}
			});
		});
	}

	window.init = function () {
		resolveParent();
		loadPurchaseTypes();
		return false;
	};

	window.ChangePurType = function () {
		resolveParent();
		if (purTypeRoot) {
			appendMatchingPurAcc();
		} else {
			loadPurchaseTypes(appendMatchingPurAcc);
		}
		return false;
	};

	window.FinalSubmit = function () {
		resolveParent();
		if (parentRoot) {
			parentRoot.setAttribute("Confirm", "Y");
		}
		returnRoot();
		window.close();
		return false;
	};

	window.addEventListener("beforeunload", returnRoot);
}(window, document));
