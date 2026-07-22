(function (window, document) {
	"use strict";

	var dataDoc = null;
	var dataRoot = null;

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

	function selectedOption(name) {
		var select = field(name);
		if (!select || select.selectedIndex < 0) {
			return { value: "", text: "" };
		}
		return {
			value: select.options[select.selectedIndex].value,
			text: select.options[select.selectedIndex].text
		};
	}

	function selectByValue(name, value) {
		var select = field(name);
		var wanted = trim(value).toLowerCase();
		if (!select || !select.options) {
			return;
		}
		for (var i = 0; i < select.options.length; i += 1) {
			if (trim(select.options[i].value).toLowerCase() === wanted) {
				select.selectedIndex = i;
				return;
			}
		}
	}

	function resolveData() {
		ensureCompat();
		var args = modalArgs();
		dataDoc = xmlDocument(args) || xmlDocument(xmlObject("RefData"));
		dataRoot = xmlRoot(args) || xmlRoot(xmlObject("RefData"));
	}

	function removeChildren(name) {
		elementChildren(dataRoot, name).forEach(function (node) {
			dataRoot.removeChild(node);
		});
	}

	function returnRoot() {
		if (!dataRoot) {
			return;
		}
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(dataRoot);
		} else {
			window["return" + "Value"] = dataRoot;
			window.returnvalue = dataRoot;
		}
	}

	function createElement(name) {
		return (dataDoc || document).createElement(name);
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
		}
		alert("Modern browser compatibility script is still loading. Please try again.");
		return null;
	}

	function applyPartySelection(outValue) {
		var root = xmlRoot(outValue);
		var action = getAttr(root, "Action");
		var query = getAttr(root, "PassQuery");
		var firstChild;
		if (!root || action === "CLOSE") {
			return;
		}
		if (action !== "Done") {
			openDialog("../../Common/PartySelection.asp?" + query, xmlObject("PartyData"), "dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", applyPartySelection);
			return;
		}
		firstChild = elementChildren(root)[0];
		if (firstChild) {
			if (field("hSupplierName")) {
				field("hSupplierName").value = getAttr(firstChild, "RetField0");
			}
			if (field("hSupplier")) {
				field("hSupplier").value = getAttr(firstChild, "RetField1");
			}
		}
	}

	window.Init = function () {
		var refNode;
		resolveData();
		if (!dataRoot || !dataRoot.hasChildNodes()) {
			return false;
		}
		setAttr(dataRoot, "Done", "N");
		selectByValue("selUsage", getAttr(dataRoot, "Usage"));
		refNode = elementChildren(dataRoot, "Ref")[0];
		if (refNode) {
			selectByValue("selIssueFor", getAttr(refNode, "Issue"));
			window.popParty();
		}
		return false;
	};

	window.FinalSubmit = function () {
		var usage = selectedOption("selUsage");
		var issueFor = selectedOption("selIssueFor");
		var partyNode;
		var refNode;
		var supplierName = field("hSupplierName") ? field("hSupplierName").value : "";
		var supplierCode = field("hSupplier") ? field("hSupplier").value : "";
		resolveData();
		if (!field("selUsage") || field("selUsage").selectedIndex < 0) {
			if (field("selUsage")) {
				field("selUsage").focus();
			}
			alert("Select Usage");
			return false;
		}
		if (!field("selIssueFor") || field("selIssueFor").selectedIndex < 0) {
			if (field("selIssueFor")) {
				field("selIssueFor").focus();
			}
			alert("Select Issue To");
			return false;
		}
		if (!dataRoot) {
			return false;
		}
		setAttr(dataRoot, "Usage", usage.value);
		setAttr(dataRoot, "UsageName", usage.text);
		setAttr(dataRoot, "Done", "Y");
		setAttr(dataRoot, "IssueTo", issueFor.value);
		setAttr(dataRoot, "IssueToName", issueFor.text);
		removeChildren("Party");
		removeChildren("Ref");
		partyNode = createElement("Party");
		setAttr(partyNode, "Name", supplierName && supplierCode ? supplierName : "");
		setAttr(partyNode, "Code", supplierName && supplierCode ? supplierCode : "");
		dataRoot.appendChild(partyNode);
		refNode = createElement("Ref");
		setAttr(refNode, "Issue", issueFor.value);
		setAttr(refNode, "IssName", issueFor.text);
		dataRoot.appendChild(refNode);
		returnRoot();
		window.close();
		return false;
	};

	window.popParty = function () {
		var issueFor = selectedOption("selIssueFor").value;
		var orgId;
		if (trim(field("selUsage") && field("selUsage").selectedIndex) === "-1") {
			alert("Select Usage");
			if (field("selUsage")) {
				field("selUsage").focus();
			}
			if (field("selIssueFor")) {
				field("selIssueFor").value = "A";
			}
			return false;
		}
		if (trim(issueFor).toLowerCase() === "party") {
			if (field("hUsage")) {
				field("hUsage").value = selectedOption("selUsage").value;
			}
			orgId = field("hUnit") ? field("hUnit").value : "";
			openDialog("../../Common/PartySelection.asp?orgID=" + encodeURIComponent(orgId), xmlObject("PartyData"), "dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", applyPartySelection);
		}
		return false;
	};

	window.addEventListener("beforeunload", returnRoot);
}(window, document));
