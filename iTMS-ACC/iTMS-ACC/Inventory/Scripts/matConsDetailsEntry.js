(function (window, document) {
	"use strict";

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function form() {
		return document.forms.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements && frm.elements[name] || document.getElementById(name) || null;
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function xmlIsland(name) {
		ensureCompat();
		return window[name] || document[name] || document.getElementById(name);
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
		return value.documentElement || value.XMLDocument && value.XMLDocument.documentElement || value._doc && value._doc.documentElement || (value.nodeType === 1 ? value : null);
	}

	function serializeXml(value) {
		var doc = xmlDocument(value);
		var root = xmlRoot(value);
		if (value && typeof value.xml === "string") {
			return value.xml;
		}
		if (doc && typeof doc.xml === "string") {
			return doc.xml;
		}
		return new XMLSerializer().serializeToString(doc || root);
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
		if (node) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function outRoot() {
		return xmlRoot(xmlIsland("OutData"));
	}

	function outDoc() {
		return xmlDocument(xmlIsland("OutData")) || outRoot() && outRoot().ownerDocument;
	}

	function openModal(url, args, features, callback) {
		ensureCompat();
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return null;
		}
		return window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	}

	function syncPost(url, xmlText) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		xhr.setRequestHeader("Content-Type", "text/xml");
		xhr.send(xmlText || "");
		return xhr;
	}

	function findItemDet(issueNo, itemCode, classCode, issueOnly) {
		var nodes = elementChildren(outRoot(), "ItemDet");
		for (var i = 0; i < nodes.length; i += 1) {
			if (getAttr(nodes[i], "IssEntryNo") !== trim(issueNo)) {
				continue;
			}
			if (issueOnly || (getAttr(nodes[i], "Item") === trim(itemCode) && getAttr(nodes[i], "Class") === trim(classCode))) {
				return nodes[i];
			}
		}
		return null;
	}

	function ensureItemDet(issueNo, itemCode, classCode, issueOnly) {
		var root = outRoot();
		var doc = outDoc();
		var node = findItemDet(issueNo, itemCode, classCode, issueOnly);
		if (node || !root || !doc) {
			return node;
		}
		node = doc.createElement("ItemDet");
		setAttr(node, "Item", issueOnly ? "" : itemCode);
		setAttr(node, "Class", issueOnly ? "" : classCode);
		setAttr(node, "IssEntryNo", issueNo);
		setAttr(node, "Remarks", "");
		setAttr(node, "Qty", "0");
		setAttr(node, "AttributeList", "");
		root.appendChild(node);
		return node;
	}

	function sumChildQuantity(itemDet, childName, attrName) {
		var total = 0;
		elementChildren(itemDet, childName).forEach(function (node) {
			total += toNumber(getAttr(node, attrName));
		});
		return total;
	}

	function updateConsumptionQty(issueNo, itemCode, classCode, rowKey, issueOnly) {
		var itemDet = findItemDet(issueNo, itemCode, classCode, issueOnly);
		var target = field("txtConsumeQtyZ" + rowKey);
		var qty = 0;
		if (itemDet) {
			qty = sumChildQuantity(itemDet, "LotDet", "QtyRet") || toNumber(getAttr(itemDet, "Qty"));
		}
		if (target) {
			target.value = qty;
		}
	}

	function updateAddDetailsQty(issueNo, itemCode, classCode, rowKey, issueOnly) {
		var itemDet = findItemDet(issueNo, itemCode, classCode, issueOnly);
		var target = field("txtConsumeQtyZ" + rowKey);
		if (itemDet && target && getAttr(itemDet, "Qty") !== "") {
			target.value = getAttr(itemDet, "Qty");
		}
	}

	function paginationNode(itemDet) {
		return elementChildren(itemDet, "Pagination")[0] || null;
	}

	function reopenForPagination(popupName, features, issueNo, itemCode, classCode, rowKey, issueOnly, afterDone) {
		var itemDet = findItemDet(issueNo, itemCode, classCode, issueOnly);
		var page = paginationNode(itemDet);
		var details = page ? getAttr(page, "Details") : "";
		var parts = details.split("```");
		if (details && parts.length === 1) {
			openModal(popupName + "?sTemp=" + details.replace(/\|/g, "&"), xmlIsland("OutData"), features, function () {
				reopenForPagination(popupName, features, issueNo, itemCode, classCode, rowKey, issueOnly, afterDone);
			});
			return;
		}
		if (parts.length > 1 && trim(parts[1]).substr(0, 2).toUpperCase() === "NO") {
			return;
		}
		if (afterDone) {
			afterDone();
		}
	}

	function receiptTemp(issueNo, itemCode, classCode, orgId, issueOnly) {
		return "F:" + issueNo + ":" + (issueOnly ? "" : itemCode) + ":" + (issueOnly ? "" : classCode) + ":" + orgId + ":Material Consumption:Consumed::" + (issueOnly ? "ISSNO" : "ITEM");
	}

	function lotTemp(issueNo, itemCode, classCode, orgId, issueOnly) {
		return "F:" + issueNo + ":" + (issueOnly ? "" : itemCode) + ":" + (issueOnly ? "" : classCode) + ":" + orgId + ":Material Consumption:Consumed:::"
			+ (issueOnly ? "ISSNO" : "ITEM");
	}

	window.ReceiptSelection = function (issueNo, itemCode, classCode, orgId, rowNo) {
		var features = "dialogHeight:580px;dialogWidth:580px;center:Yes;help:No;resizable:No;status:No";
		ensureItemDet(issueNo, itemCode, classCode, false);
		openModal("MatConRcptSelPop.asp?sTemp=" + encodeURIComponent(receiptTemp(issueNo, itemCode, classCode, orgId, false)), xmlIsland("OutData"), features, function () {
			reopenForPagination("MatConRcptSelPop.asp", features, issueNo, itemCode, classCode, rowNo, false);
		});
		return false;
	};

	window.ReceiptSelectionIss = function (issueNo, orgId) {
		var features = "dialogHeight:580px;dialogWidth:580px;center:Yes;help:No;resizable:No;status:No";
		ensureItemDet(issueNo, "", "", true);
		openModal("MatConRcptSelPop.asp?sTemp=" + encodeURIComponent(receiptTemp(issueNo, "", "", orgId, true)), xmlIsland("OutData"), features, function () {
			reopenForPagination("MatConRcptSelPop.asp", features, issueNo, "", "", issueNo, true);
		});
		return false;
	};

	window.PackSelection = function (issueNo, itemCode, classCode, orgId, rowNo) {
		ensureItemDet(issueNo, itemCode, classCode, false);
		openModal("MatConLotSerPop.asp?sTemp=" + encodeURIComponent(lotTemp(issueNo, itemCode, classCode, orgId, false)), xmlIsland("OutData"), "dialogHeight:320px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:Yes", function () {
			updateConsumptionQty(issueNo, itemCode, classCode, rowNo, false);
		});
		return false;
	};

	window.PackSelectionIss = function (issueNo, orgId) {
		ensureItemDet(issueNo, "", "", true);
		openModal("MatConLotSerPop.asp?sTemp=" + encodeURIComponent(lotTemp(issueNo, "", "", orgId, true)), xmlIsland("OutData"), "dialogHeight:320px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:Yes", function () {
			updateConsumptionQty(issueNo, "", "", issueNo, true);
		});
		return false;
	};

	window.AddDetails = function (issueNo, itemCode, classCode, orgId, rowNo, issuedForCode, itemEntryNo, attributeList) {
		var qtyField = field("txtConsumeQtyZ" + rowNo);
		var quantity = qtyField && trim(qtyField.value) !== "" ? toNumber(qtyField.value) : 0;
		var tempValues;
		ensureItemDet(issueNo, itemCode, classCode, false);
		if (quantity <= 0) {
			alert("Select the Consumption Quantity");
			return false;
		}
		tempValues = classCode + "|" + itemCode + "|" + orgId + "|" + (issuedForCode || "") + "|" + quantity + "|" + (itemEntryNo || "") + "|" + (attributeList || "") + "|" + issueNo;
		openModal("MatConAddEntryDetails.asp?sTemp=" + encodeURIComponent(tempValues), xmlIsland("OutData"), "dialogHeight:320px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:Yes", function () {
			updateAddDetailsQty(issueNo, itemCode, classCode, rowNo, false);
		});
		return false;
	};

	window.AddDetailsIss = function (issueNo) {
		alert("Additional details are available from the item-wise consumption rows.");
		return false;
	};

	window.CheckSubmit = function () {
		var root = outRoot();
		var nodes = elementChildren(root, "ItemDet");
		var issueNo = field("hIssueNo") ? field("hIssueNo").value : "";
		var remarks;
		if (!nodes.length) {
			alert("No Consumption are Selected");
			return false;
		}
		nodes.forEach(function (node) {
			if (trim(issueNo)) {
				remarks = field("txtRemarksZ" + getAttr(node, "IssEntryNo"));
			} else {
				remarks = field("txtRemarksZ" + getAttr(node, "Item") + "Z" + getAttr(node, "Class") + "Z" + getAttr(node, "IssEntryNo"));
			}
			setAttr(node, "Remarks", remarks ? remarks.value : "");
		});
		syncPost("XMLSave.asp?SessionFlag=true&Name=MatConsumption", serializeXml(xmlIsland("OutData")));
		form().action = "MatConDetInsert.asp";
		form().submit();
		return false;
	};
}(window, document));
