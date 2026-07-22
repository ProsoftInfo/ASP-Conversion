(function (window, document) {
	"use strict";

	var objTemp = null;
	var rootNode = null;
	var iClass = "";
	var iItem = "";
	var iQty = 0;
	var iEntNo = "";

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
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

	function modalArgs() {
		ensureCompat();
		return window["dialog" + "Arguments"] || null;
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function field(name) {
		var form = document.forms.formname || document.forms[0];
		return form && form.elements ? form.elements[name] : null;
	}

	function checkNumbers(value) {
		return /^[0-9.]+$/.test(String(value || ""));
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

	function attrAt(node, index) {
		var attr = node && node.attributes && node.attributes.item(index);
		return trim(attr ? attr.nodeValue || attr.value || "" : "");
	}

	function getHeaderNode() {
		var headers = elementChildren(rootNode);
		for (var i = 0; i < headers.length; i += 1) {
			if (attrAt(headers[i], 0) === trim(iEntNo) && attrAt(headers[i], 1) === trim(iItem) && attrAt(headers[i], 2) === trim(iClass)) {
				return headers[i];
			}
		}
		return null;
	}

	function setReturnValue() {
		if (!rootNode) {
			return;
		}
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(rootNode);
		} else {
			window["return" + "Value"] = rootNode;
			window.returnvalue = rootNode;
		}
	}

	function closeWithReturn() {
		setReturnValue();
		window.close();
	}

	window.fnInit = function (sItem, sClass, sQty, sEntNo) {
		var headerNode;
		var stNodes;
		var hiCtr;
		var q;
		var i;

		iClass = sClass;
		iItem = sItem;
		iQty = toNumber(sQty);
		iEntNo = sEntNo;
		objTemp = modalArgs();
		rootNode = xmlRoot(objTemp);
		headerNode = getHeaderNode();
		hiCtr = toNumber(field("hiCtr") && field("hiCtr").value);

		if (!hiCtr || !headerNode) {
			return;
		}

		stNodes = elementChildren(headerNode, "STDETAILS");
		if (stNodes.length) {
			for (i = 0; i < stNodes.length; i += 1) {
				q = field("txtST" + (i + 1));
				if (q) {
					q.value = attrAt(stNodes[i], 1);
				}
			}
		} else {
			for (i = 1; i <= hiCtr; i += 1) {
				q = field("txtST" + i);
				if (q) {
					q.value = "0";
				}
			}
		}
	};

	window.CheckSubmit = function () {
		var doc = xmlDocument(objTemp);
		var headerNode = getHeaderNode();
		var ictr = toNumber(field("hiCtr") && field("hiCtr").value);
		var iQtyTot = 0;
		var objQ;
		var objh;
		var sQty;
		var stNodes;
		var newElem;
		var hMrsNo;
		var i;

		if (!ictr || !doc || !headerNode) {
			return;
		}

		for (i = 1; i <= ictr; i += 1) {
			objQ = field("txtST" + i);
			objh = field("hST" + i);

			if (!objQ || trim(objQ.value) === "") {
				alert("Enter Quantity");
				if (objQ) {
					objQ.select();
				}
				return;
			}
			if (!checkNumbers(objQ.value)) {
				alert("Enter Numerals Only");
				objQ.select();
				return;
			}
			sQty = objQ.value;
			if (toNumber(sQty) > toNumber(objh && objh.value)) {
				alert("Transfer Quantity should be equal to or less than Stock Quantity " + (objh ? objh.value : ""));
				objQ.select();
				return;
			}
			iQtyTot += toNumber(objQ.value);
		}

		stNodes = elementChildren(headerNode, "STDETAILS");
		for (i = 0; i < stNodes.length; i += 1) {
			headerNode.removeChild(stNodes[i]);
		}

		if (iQtyTot !== 0) {
			for (i = 1; i <= ictr; i += 1) {
				newElem = doc.createElement("STDETAILS");
				newElem.setAttribute("ORGID", trim(field("hOrgID" + i) && field("hOrgID" + i).value));
				newElem.setAttribute("QTY", trim(field("txtST" + i) && field("txtST" + i).value));
				headerNode.appendChild(newElem);
			}
		}

		hMrsNo = field("hMrsNo");
		if (hMrsNo && trim(hMrsNo.value) !== "" && iQtyTot > iQty) {
			alert("Total Transfer Quantity should be equal to or less than Quantity Pending " + iQty);
			return;
		}

		headerNode.setAttribute("TRAQTY", iQtyTot);
		closeWithReturn();
	};

	window.window_onunload = setReturnValue;
	window.addEventListener("beforeunload", setReturnValue);
}(window, document));
