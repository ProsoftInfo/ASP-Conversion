(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		if (!frm) {
			return null;
		}
		return frm.elements[name] || frm.elements[name.toLowerCase()] || frm.elements[name.toUpperCase()] || null;
	}

	function setValue(name, value) {
		var element = field(name);
		if (element) {
			element.value = value == null ? "" : String(value);
		}
	}

	function xmlRoot(value, fallback) {
		var candidate = value || fallback;
		if (!candidate) {
			return null;
		}
		if (candidate.documentElement) {
			return candidate.documentElement;
		}
		if (candidate.XMLDocument && candidate.XMLDocument.documentElement) {
			return candidate.XMLDocument.documentElement;
		}
		if (candidate._doc && candidate._doc.documentElement) {
			return candidate._doc.documentElement;
		}
		if (candidate.nodeType === 1) {
			return candidate;
		}
		return null;
	}

	function childElements(node, name) {
		var wanted = name && String(name).toLowerCase();
		return Array.prototype.filter.call(node && node.childNodes || [], function (child) {
			return child.nodeType === 1 && (!wanted || String(child.nodeName).toLowerCase() === wanted);
		});
	}

	function getAttr(node, name) {
		return node && node.getAttribute ? trim(node.getAttribute(name)) : "";
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function getXmlText(xmlObject) {
		var target = xmlObject && (xmlObject.XMLDocument || xmlObject._doc || xmlObject);
		if (!target) {
			return "";
		}
		if (typeof target.xml === "string") {
			return target.xml;
		}
		return new XMLSerializer().serializeToString(target);
	}

	function openDialog(url, args, features, callback) {
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return;
		}
		window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	}

	function getDateValue(name) {
		var control = field(name);
		if (!control) {
			return "";
		}
		if (typeof control.GetDate === "function") {
			return control.GetDate();
		}
		if (typeof control.getDate === "function") {
			return control.getDate();
		}
		if (window.ITMSModernCompat && window.ITMSModernCompat.toDisplayDate) {
			return window.ITMSModernCompat.toDisplayDate(control.value);
		}
		return trim(control.value);
	}

	function findItemNode(classCode, itemCode, itemEntryNo) {
		var root = window.ItemData && window.ItemData.documentElement;
		var matches = childElements(root, "ITM");
		var i;
		for (i = 0; i < matches.length; i += 1) {
			if (
				getAttr(matches[i], "CLACODE") === trim(classCode) &&
				getAttr(matches[i], "ITMCODE") === trim(itemCode) &&
				getAttr(matches[i], "ItemEntNo") === trim(itemEntryNo)
			) {
				return matches[i];
			}
		}
		return null;
	}

	function pickedQuantity(classCode, itemCode, itemEntryNo) {
		var itemNode = findItemNode(classCode, itemCode, itemEntryNo);
		var pickDet = childElements(itemNode, "PickDet")[0];
		return pickDet ? getAttr(pickDet, "TOT") : "0";
	}

	function appendMatchingChildren(targetRoot, sourceRoot, nodeName) {
		childElements(sourceRoot, nodeName).forEach(function (node) {
			targetRoot.appendChild(targetRoot.ownerDocument.importNode ? targetRoot.ownerDocument.importNode(node, true) : node.cloneNode(true));
		});
	}

	function updateIssueQuantities(root) {
		childElements(root, "ITM").forEach(function (node) {
			var classCode = getAttr(node, "CLACODE");
			var itemCode = getAttr(node, "ITMCODE");
			var pickNo = getAttr(node, "PICKNO");
			var itemEntryNo = getAttr(node, "ItemEntNo");
			var input = field("txtQtyPA" + classCode + "A" + itemCode + "A" + pickNo + "A" + itemEntryNo);
			setAttr(node, "ISSQTY", input ? input.value : "");
		});
	}

	function saveXmlAndSubmit(tempValues) {
		var xhr = new XMLHttpRequest();
		var scheduleNo = trim(field("hScheduleNo") && field("hScheduleNo").value);
		xhr.open("POST", "XMLSave.asp?SessionFlag=true&Name=IssuePick_", true);
		xhr.setRequestHeader("Content-Type", "text/xml; charset=UTF-8");
		xhr.onreadystatechange = function () {
			if (xhr.readyState !== 4) {
				return;
			}
			form().action = "mrsIssuePickInsert.asp?TEMPVALUES=" + encodeURIComponent(tempValues) + "&ScheduleNo=" + encodeURIComponent(scheduleNo);
			form().submit();
		};
		xhr.send(getXmlText(window.ItemData));
	}

	window.GetPick = function (classCode, itemCode, quantity, issueEntryNo, attributeId, itemEntryNo) {
		var tempValues = [classCode, itemCode, quantity, issueEntryNo, attributeId, itemEntryNo].join("|");
		openDialog("mrsIssuePickDetailsEntry.asp?sTemp=" + encodeURIComponent(tempValues), window.ItemData, "dialogHeight:420px;dialogWidth:680px;center:Yes;help:No;resizable:No;status:No", function (result) {
			var target = field("txtQtyPA" + classCode + "A" + itemCode + "A" + issueEntryNo + "A" + itemEntryNo);
			xmlRoot(result, window.ItemData);
			if (target) {
				target.value = pickedQuantity(classCode, itemCode, itemEntryNo);
			}
		});
	};

	window.CheckSubmit = function () {
		var root = window.ItemData && window.ItemData.documentElement;
		var issueDate = getDateValue("ctlPickDate");
		var receivedBy = trim(field("txtRecdBy") && field("txtRecdBy").value);
		var tempValues = issueDate + ":" + receivedBy;
		if (!root) {
			return;
		}
		setValue("IssDate", issueDate);
		setValue("RecBy", receivedBy);
		updateIssueQuantities(root);
		root.setAttribute("POConfirm", "N");
		root.setAttribute("SInvConfirm", "N");
		root.setAttribute("Invoice", "A");
		root.setAttribute("GPConfirm", "N");
		root.setAttribute("ProConfirm", "N");
		root.setAttribute("PickDate", issueDate);
		saveXmlAndSubmit(tempValues);
	};

	window.__itmsAppendMatchingChildren = appendMatchingChildren;
})(window, document);
