(function (window, document) {
	"use strict";

	var dialogDoc = normalizeXmlDocument(modalArgs());

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function modalArgs() {
		ensureCompat();
		return window["dialog" + "Arguments"] || null;
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/\u00a0/g, " ").replace(/^\s+|\s+$/g, "");
	}

	function upper(value) {
		return trim(value).toUpperCase();
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

	function setSpanText(id, value) {
		var element = document.getElementById(id) || window[id];
		if (element) {
			element.textContent = value || "";
		}
	}

	function normalizeXmlDocument(value) {
		if (value && value.XMLDocument) {
			return value.XMLDocument;
		}
		if (value && value._doc) {
			return value._doc;
		}
		if (value && value.nodeType === 9) {
			return value;
		}
		if (value && value.documentElement && value.createElement) {
			return value;
		}
		if (value && value.nodeType === 1 && value.ownerDocument) {
			return value.ownerDocument;
		}
		return document.implementation.createDocument("", "Root", null);
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

	function childElements(node) {
		return Array.prototype.filter.call(node && node.childNodes || [], function (child) {
			return child.nodeType === 1;
		});
	}

	function getAttr(node, name) {
		return node && node.getAttribute ? trim(node.getAttribute(name)) : "";
	}

	function openDialog(url, args, features, callback) {
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return;
		}
		window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	}

	function popupSize() {
		var text = typeof window.GetWindowSizeForPopup === "function" ? window.GetWindowSizeForPopup("1") : "";
		var parts = text.split(":");
		return {
			program: parts[0] || "ItemSelectRelPartyCommon.asp",
			height: parts[1] || "500",
			width: parts[2] || "850"
		};
	}

	function processItemSelection(root) {
		var names = [];
		var values = [];
		var attrs = [];
		if (!root || !root.hasChildNodes()) {
			return;
		}
		childElements(root).forEach(function (node) {
			var itemCode = getAttr(node, "ItemCode");
			var classCode = getAttr(node, "ClassCode");
			names.push(getAttr(node, "ItemName"));
			values.push(itemCode + ":" + classCode);
			attrs.push(getAttr(node, "AttributeList"));
		});
		if (names.length) {
			setSpanText("SpnMaterialToBeReceived", names.join(","));
			setValue("cmbMatRecdAs", values.join(","));
		}
	}

	function continueItemDialog(result, xmlIsland, size) {
		var action = upper(getAttr(result, "Action"));
		var passQuery = getAttr(result, "PassQuery");
		if (action && action !== "DONE" && action !== "CLOSE") {
			openDialog("../../Common/" + size.program + "?" + passQuery, xmlIsland, "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", function (nextResult) {
				continueItemDialog(nextResult, xmlIsland, size);
			});
			return;
		}
		if (action !== "CLOSE") {
			processItemSelection(xmlRoot(result, xmlIsland));
		}
	}

	function selectedText(select) {
		return select && select.options && select.selectedIndex >= 0 ? trim(select.options[select.selectedIndex].text) : "";
	}

	function selectByValue(name, value) {
		var control = field(name);
		var i;
		if (!control || !control.options) {
			return;
		}
		for (i = 0; i < control.options.length; i += 1) {
			if (trim(control.options[i].value) === trim(value)) {
				control.selectedIndex = i;
				break;
			}
		}
	}

	function checkNumbers(value) {
		return /^[0-9]+$/.test(trim(value));
	}

	function focusField(name, selectText) {
		var element = field(name);
		if (!element) {
			return;
		}
		if (selectText && element.select) {
			element.select();
		} else if (element.focus) {
			element.focus();
		}
	}

	function removeNamedChildren(root, nodeName) {
		childElements(root).forEach(function (node) {
			if (node.nodeName === nodeName) {
				root.removeChild(node);
			}
		});
	}

	function publishReturnValue() {
		var root = dialogDoc && dialogDoc.documentElement;
		if (!root) {
			return;
		}
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(root);
		} else {
			window["return" + "Value"] = root;
			window.returnvalue = root;
		}
	}

	function closeWithReturnValue() {
		publishReturnValue();
		window.close();
	}

	window.checkNumbers = checkNumbers;

	window.SelectItem = function () {
		var unit = trim(field("hUnit") && field("hUnit").value);
		var size = popupSize();
		var url = "../../Common/" + size.program +
			"?orgID=" + encodeURIComponent(unit) +
			"&sIType=&Stock=Y&hSelectMode=M&Flag=&hDispButt=Y&PartyType=CR&CallFrom=PUR";
		openDialog(url, window.Data, "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", function (result) {
			continueItemDialog(result, window.Data, size);
		});
	};

	window.saveXML = function () {
		var root = dialogDoc.documentElement;
		var processControl = field("cmbSCProcess");
		var materialReceivedAs = trim(field("cmbMatRecdAs") && field("cmbMatRecdAs").value);
		var labourCharge = trim(field("txtLabourCharge") && field("txtLabourCharge").value);
		var hardWaste = trim(field("txtHardWaste") && field("txtHardWaste").value) || "0";
		var invWaste = trim(field("txtInvWaste") && field("txtInvWaste").value) || "0";
		var processNode;
		var materialValues;
		var materialNames;

		if (!checkNumbers(labourCharge)) {
			alert("Enter Only Numberic");
			focusField("txtLabourCharge", true);
			return;
		}
		if (trim(processControl && processControl.value) === "0") {
			alert("Select Subcontracting Process");
			focusField("cmbSCProcess");
			return;
		}
		if (!trim(field("txtInstruct") && field("txtInstruct").value)) {
			alert("Enter Instruction");
			focusField("txtInstruct");
			return;
		}
		if (materialReceivedAs === "0") {
			alert("Select Material Received As");
			return;
		}

		removeNamedChildren(root, "SubContract");

		processNode = dialogDoc.createElement("SubContract");
		processNode.setAttribute("SCProcess", trim(processControl && processControl.value));
		processNode.setAttribute("Instruct", field("txtInstruct").value);
		processNode.setAttribute("LabourCharge", labourCharge);
		processNode.setAttribute("Currency", trim(field("cmbCurrency") && field("cmbCurrency").value));
		processNode.setAttribute("HardWaste", hardWaste);
		processNode.setAttribute("InvWaste", invWaste);
		processNode.setAttribute("ProcessName", selectedText(processControl));
		root.appendChild(processNode);

		if (materialReceivedAs && materialReceivedAs !== "0") {
			materialValues = materialReceivedAs.split(",");
			materialNames = trim((document.getElementById("SpnMaterialToBeReceived") || {}).textContent).split(",");
			materialValues.forEach(function (value, index) {
				var itemParts = value.split(":");
				var detailsNode = dialogDoc.createElement("Details");
				detailsNode.setAttribute("MatRecdAsItem", trim(itemParts[0]));
				detailsNode.setAttribute("MatRecdAsCode", trim(itemParts[1]));
				detailsNode.setAttribute("MatRecdAsDescr", trim(materialNames[index]));
				processNode.appendChild(detailsNode);
			});
		}

		closeWithReturnValue();
	};

	window.Init = function () {
		var root = dialogDoc && dialogDoc.documentElement;
		var names = [];
		var values = [];
		if (!root || !root.hasChildNodes()) {
			return;
		}
		childElements(root).forEach(function (node) {
			if (node.nodeName !== "SubContract") {
				return;
			}
			selectByValue("cmbSCProcess", getAttr(node, "SCProcess"));
			selectByValue("cmbCurrency", getAttr(node, "Currency"));
			setValue("txtInstruct", getAttr(node, "Instruct"));
			setValue("txtLabourCharge", getAttr(node, "LabourCharge"));
			setValue("txtHardWaste", getAttr(node, "HardWaste"));
			setValue("txtInvWaste", getAttr(node, "InvWaste"));
			childElements(node).forEach(function (detail) {
				if (detail.nodeName === "Details") {
					values.push(getAttr(detail, "MatRecdAsItem") + ":" + getAttr(detail, "MatRecdAsCode"));
					names.push(getAttr(detail, "MatRecdAsDescr"));
				}
			});
		});
		if (names.length) {
			setSpanText("SpnMaterialToBeReceived", names.join(","));
			setValue("cmbMatRecdAs", values.join(","));
		}
	};

	window.addEventListener("beforeunload", publishReturnValue);
})(window, document);
