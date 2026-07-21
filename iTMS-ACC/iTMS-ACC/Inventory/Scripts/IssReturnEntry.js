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
		return frm && frm.elements ? frm.elements[name] || null : null;
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

	function serializeXml(value) {
		var target = value && value.nodeType === 9 ? value : xmlDocument(value);
		if (!target) {
			return "";
		}
		return new XMLSerializer().serializeToString(target);
	}

	function openModal(url, args, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args, features, callback);
		}
		if (typeof window.showModalDialog === "function") {
			if (callback) {
				callback(window.showModalDialog(url, args, features));
			}
			return null;
		}
		window.open(url, "_blank");
		return null;
	}

	function outRoot() {
		return xmlRoot(xmlIsland("OutData"));
	}

	function updateReturnQty(itemCode, classCode, attrId, rowNo) {
		var nodes = elementChildren(outRoot(), "ITEM");
		var qty = field("txtConsumeQtyZ" + rowNo);
		for (var i = 0; i < nodes.length; i += 1) {
			if (getAttr(nodes[i], "ITMCODE") === trim(itemCode) &&
					getAttr(nodes[i], "CLACODE") === trim(classCode) &&
					getAttr(nodes[i], "ATTID") === trim(attrId)) {
				if (qty) {
					qty.value = getAttr(nodes[i], "QTY");
				}
				return;
			}
		}
	}

	function updateAddDetailsQty(issueNo, itemCode, classCode, rowNo) {
		var nodes = elementChildren(outRoot());
		var qty = field("txtConsumeQtyZ" + rowNo);
		for (var i = 0; i < nodes.length; i += 1) {
			if (getAttr(nodes[i], "Item") === trim(itemCode) &&
					getAttr(nodes[i], "Class") === trim(classCode) &&
					getAttr(nodes[i], "IssEntryNo") === trim(issueNo)) {
				if (qty) {
					qty.value = getAttr(nodes[i], "Qty");
				}
				return;
			}
		}
	}

	function postIssueReturnXml() {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", "XMLSave.asp?SessionFlag=true&Name=IssueReturn", false);
		try {
			xhr.setRequestHeader("Content-Type", "text/xml");
		} catch (ignoreHeader) {}
		xhr.send(serializeXml(xmlIsland("OutData")));
		return xhr;
	}

	window.PackSelection = function (issueNo, itemCode, classCode, orgId, rowNo, attributeList, itemEntryNo) {
		var tempValues = "F:" + issueNo + ":" + itemCode + ":" + classCode + ":" + orgId + ":Material Return:Returned:" + attributeList + ":" + itemEntryNo;
		openModal("IssRetLotSerPop.asp?sTemp=" + encodeURIComponent(tempValues), xmlIsland("OutData"), "dialogHeight:320px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:Yes", function () {
			updateReturnQty(itemCode, classCode, attributeList, rowNo);
		});
		return false;
	};

	window.AddDetails = function (issueNo, itemCode, classCode, orgId, rowNo, itemEntryNo, attributeList, issuedForCode) {
		var qty = field("txtConsumeQtyZ" + rowNo);
		var quantity = qty && trim(qty.value) !== "" ? toNumber(qty.value) : 0;
		var tempValues;
		if (quantity <= 0) {
			alert("Select the Consumption Quantity");
			return false;
		}
		tempValues = classCode + "|" + itemCode + "|" + orgId + "|" + issuedForCode + "|" + quantity + "|" + itemEntryNo + "|" + attributeList + "|" + issueNo;
		openModal("MatConAddEntryDetails.asp?sTemp=" + encodeURIComponent(tempValues), xmlIsland("OutData"), "dialogHeight:320px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:Yes", function () {
			updateAddDetailsQty(issueNo, itemCode, classCode, rowNo);
		});
		return false;
	};

	window.CheckSubmit = function () {
		var root = outRoot();
		var frm = form();
		var issue = field("hIssueEntryNo");
		if (root && elementChildren(root).length) {
			postIssueReturnXml();
			frm.action = "IssReturnDetInsert.asp?IssueNo=" + encodeURIComponent(issue ? issue.value : "");
			frm.submit();
			return false;
		}
		alert("No Return are Selected");
		return false;
	};
}(window, document));
