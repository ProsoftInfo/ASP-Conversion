(function (window, document) {
	"use strict";

	var objTemp = null;
	var rootNode = null;
	var iClass = "";
	var iItem = "";

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

	function textOf(id) {
		var item = document.getElementById(id) || window[id] || null;
		return trim(item ? item.innerText || item.textContent || "" : "");
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

	function attr(node, name) {
		return node && node.getAttribute ? trim(node.getAttribute(name)) : "";
	}

	function getHeaderNode() {
		var items = rootNode ? rootNode.getElementsByTagName("ITEM") : [];
		for (var i = 0; i < items.length; i += 1) {
			if (attr(items[i], "CLACODE") === trim(iClass) && attr(items[i], "ITMCODE") === trim(iItem)) {
				return items[i];
			}
		}
		return null;
	}

	function getPackNodes(headerNode) {
		var groups = elementChildren(headerNode, "PackingDet");
		return groups.length ? elementChildren(groups[0], "PCode") : [];
	}

	function setReturnValue() {
		if (!rootNode) {
			return;
		}
		window.returnValue = rootNode;
		window.returnvalue = rootNode;
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(rootNode);
		}
	}

	function closeWithReturn() {
		setReturnValue();
		window.close();
	}

	window.fnInit = function (sOrg, sClass, sItm, sUsag) {
		var packNodes;
		var oQty;
		var oMC;
		var i;

		objTemp = window.dialogArguments;
		rootNode = xmlRoot(objTemp);
		iClass = sClass;
		iItem = sItm;

		packNodes = getPackNodes(getHeaderNode());
		for (i = 0; i < packNodes.length; i += 1) {
			oQty = field("txtQty" + (i + 1));
			oMC = field("hPONO" + (i + 1));
			if (oQty && oMC && trim(oMC.value) === attr(packNodes[i], "PONO")) {
				oQty.value = attr(packNodes[i], "QTY");
			}
		}
	};

	window.CheckSubmit = function () {
		var doc = xmlDocument(objTemp);
		var headerNode = getHeaderNode();
		var packNodes = getPackNodes(headerNode);
		var iCtr = toNumber(field("hiCtr") && field("hiCtr").value);
		var iTotQty = 0;
		var i;
		var oQty;
		var oMC;
		var newElem1;
		var newElem2;

		if (!iCtr || !doc || !headerNode) {
			return;
		}

		for (i = 1; i <= iCtr; i += 1) {
			oQty = field("txtQty" + i);
			if (!oQty || trim(oQty.value) === "") {
				alert("Enter Quantity");
				if (oQty) {
					oQty.select();
				}
				return;
			}
			if (toNumber(oQty.value) === 0) {
				alert("Quantity cannot be ZERO");
				oQty.select();
				return;
			}
			iTotQty += toNumber(oQty.value);
		}

		if (iTotQty !== toNumber(textOf("idQty"))) {
			alert("Quantity breakup should be equal to Quantity Issue");
			return;
		}

		if (packNodes.length > 0) {
			for (i = 0; i < packNodes.length; i += 1) {
				oQty = field("txtQty" + (i + 1));
				oMC = field("hPONO" + (i + 1));
				if (oQty && oMC && trim(oMC.value) === attr(packNodes[i], "PONO")) {
					packNodes[i].setAttribute("PONO", trim(oMC.value));
					packNodes[i].setAttribute("QTY", trim(oQty.value));
				}
			}
		} else {
			newElem1 = doc.createElement("PackingDet");
			for (i = 1; i <= iCtr; i += 1) {
				oQty = field("txtQty" + i);
				oMC = field("hPONO" + i);
				newElem2 = doc.createElement("PCode");
				newElem2.setAttribute("PONO", trim(oMC && oMC.value));
				newElem2.setAttribute("QTY", trim(oQty && oQty.value));
				newElem1.appendChild(newElem2);
			}
			headerNode.appendChild(newElem1);
		}

		closeWithReturn();
	};

	window.window_onunload = setReturnValue;
	window.addEventListener("beforeunload", setReturnValue);
}(window, document));
