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

	function processItemSelection(root, spanId, hiddenName, attrHiddenName) {
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
			setSpanText(spanId, names.join(","));
			setValue(hiddenName, values.join(","));
			if (attrHiddenName) {
				setValue(attrHiddenName, attrs.join(","));
			}
		}
	}

	function continueItemDialog(result, xmlIsland, size, processor) {
		var action = upper(getAttr(result, "Action"));
		var passQuery = getAttr(result, "PassQuery");
		if (action && action !== "DONE" && action !== "CLOSE") {
			openDialog("../../Common/" + size.program + "?" + passQuery, xmlIsland, "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", function (nextResult) {
				continueItemDialog(nextResult, xmlIsland, size, processor);
			});
			return;
		}
		if (action !== "CLOSE") {
			processor(xmlRoot(result, xmlIsland));
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

	function setRadioValue(name, value) {
		var control = field(name);
		var i;
		if (!control) {
			return;
		}
		if (typeof control.length === "number" && !control.tagName) {
			for (i = 0; i < control.length; i += 1) {
				control[i].checked = trim(control[i].value) === trim(value);
			}
			return;
		}
		control.checked = trim(control.value) === trim(value);
	}

	function getRadioValue(name) {
		var control = field(name);
		var i;
		if (!control) {
			return "";
		}
		if (typeof control.length === "number" && !control.tagName) {
			for (i = 0; i < control.length; i += 1) {
				if (control[i].checked) {
					return control[i].value;
				}
			}
			return "";
		}
		return control.checked ? control.value : "";
	}

	function selectedText(select) {
		return select && select.options && select.selectedIndex >= 0 ? trim(select.options[select.selectedIndex].text) : "";
	}

	function removeMatchingPrimary(root, itemCode, classCode, entryNo) {
		childElements(root).forEach(function (node) {
			if (
				node.nodeName === "PRIMARYADDITIONALDET" &&
				getAttr(node, "PItemCode") === trim(itemCode) &&
				getAttr(node, "PClassCode") === trim(classCode) &&
				getAttr(node, "EntryNo") === trim(entryNo)
			) {
				root.removeChild(node);
			}
		});
	}

	function attributeId(value) {
		var first = trim(value).split(":")[0];
		var parts;
		if (!first) {
			return "";
		}
		parts = first.split("#");
		return trim(parts.length === 2 ? parts[1] : parts[0]) === "0" ? "" : trim(parts.length === 2 ? parts[1] : parts[0]);
	}

	function createPrimaryAdditional(attrs) {
		var node = dialogDoc.createElement("PRIMARYADDITIONALDET");
		Object.keys(attrs).forEach(function (name) {
			node.setAttribute(name, attrs[name] == null ? "" : String(attrs[name]));
		});
		return node;
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
			"&sIType=&Stock=Y&hSelectMode=S&Flag=&hDispButt=Y&PartyType=CR&CallFrom=PUR";
		openDialog(url, window.Data, "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", function (result) {
			continueItemDialog(result, window.Data, size, function (root) {
				processItemSelection(root, "SpnMaterialToBeReceived", "cmbMatRecdAs", "hAttribute");
			});
		});
	};

	window.SelectAddItem = function () {
		var unit = trim(field("hUnit") && field("hUnit").value);
		var size = popupSize();
		var url = "../../Common/" + size.program +
			"?orgID=" + encodeURIComponent(unit) +
			"&sIType=&Stock=Y&hSelectMode=M&Flag=&hDispButt=Y&PartyType=CR&CallFrom=PUR";
		openDialog(url, window.ItemAddData, "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", function (result) {
			continueItemDialog(result, window.ItemAddData, size, function (root) {
				processItemSelection(root, "SpnAdditionalMaterials", "hAddMatAs", "hAddAttribute");
			});
		});
	};

	window.saveXML = function () {
		var root = dialogDoc.documentElement;
		var materialType = trim(getRadioValue("radType"));
		var itemCode = trim(field("hItemCode") && field("hItemCode").value);
		var classCode = trim(field("hClassCode") && field("hClassCode").value);
		var entryNo = trim(field("hEntryNo") && field("hEntryNo").value);
		var processControl = field("cmbSCProcess");
		var processId = "";
		var materialReceivedAs = "";
		var materialNames = [];
		var attributeList = [];
		var values;
		var names;
		var attrs;

		if (materialType === "P") {
			if (!checkNumbers(field("txtLabourCharge") && field("txtLabourCharge").value)) {
				alert("Enter Only Numberic");
				focusField("txtLabourCharge", true);
				return;
			}
			if (trim(field("hOrderFor") && field("hOrderFor").value) === "C") {
				if (trim(processControl && processControl.value) === "0") {
					alert("Select Subcontracting Process");
					focusField("cmbSCProcess");
					return;
				}
				if (trim(field("cmbMatRecdAs") && field("cmbMatRecdAs").value) === "0") {
					alert("Select Material Received As");
					return;
				}
				processId = trim(processControl && processControl.value);
				materialReceivedAs = trim(field("cmbMatRecdAs") && field("cmbMatRecdAs").value);
				materialNames = trim((document.getElementById("SpnMaterialToBeReceived") || {}).textContent).split(",");
				attributeList = trim(field("hAttribute") && field("hAttribute").value).split(",");
			}
			if (!trim(field("txtInstruct") && field("txtInstruct").value)) {
				alert("Enter Instruction");
				focusField("txtInstruct");
				return;
			}
		}

		removeMatchingPrimary(root, itemCode, classCode, entryNo);

		if (materialType === "P") {
			if (materialReceivedAs) {
				values = materialReceivedAs.split(",");
				values.forEach(function (value, index) {
					var itemParts = value.split(":");
					root.appendChild(createPrimaryAdditional({
						MatRecdAT: "",
						SCProcess: processId,
						MatRecdAsItem: trim(itemParts[0]),
						MatRecdAsCode: trim(itemParts[1]),
						Instruct: field("txtInstruct").value,
						LabourCharge: trim(field("txtLabourCharge") && field("txtLabourCharge").value),
						Currency: trim(field("cmbCurrency") && field("cmbCurrency").value),
						MatRecdAsDescr: trim(materialNames[index]),
						MatRecdAsItemType: "",
						AttributeList: attributeId(attributeList[index]),
						PItemCode: itemCode,
						PClassCode: classCode,
						MatType: materialType,
						EntryNo: entryNo
					}));
				});
			}
		} else {
			root.appendChild(createPrimaryAdditional({
				MatRecdAT: "",
				SCProcess: "",
				MatRecdAsItem: "",
				MatRecdAsCode: "",
				Instruct: "",
				LabourCharge: "",
				Currency: "",
				MatRecdAsDescr: "",
				MatRecdAsItemType: "",
				AttributeList: "",
				PItemCode: itemCode,
				PClassCode: classCode,
				MatType: materialType,
				EntryNo: entryNo
			}));
		}

		closeWithReturnValue();
	};

	window.ShowAdd = function () {
		var panel = document.getElementById("tblAddDet") || window.tblAddDet;
		if (panel) {
			panel.style.display = trim(getRadioValue("radType")) === "P" ? "block" : "none";
		}
	};

	window.Init = function () {
		var root = dialogDoc && dialogDoc.documentElement;
		var itemCode = trim(field("hItemCode") && field("hItemCode").value);
		var classCode = trim(field("hClassCode") && field("hClassCode").value);
		var entryNo = trim(field("hEntryNo") && field("hEntryNo").value);
		var matches = [];
		var names = [];
		var values = [];
		var attrs = [];
		var first;
		if (!root || !root.hasChildNodes()) {
			window.ShowAdd();
			return;
		}
		childElements(root).forEach(function (node) {
			if (
				node.nodeName === "PRIMARYADDITIONALDET" &&
				getAttr(node, "PItemCode") === itemCode &&
				getAttr(node, "PClassCode") === classCode &&
				getAttr(node, "EntryNo") === entryNo
			) {
				matches.push(node);
			}
		});
		if (!matches.length) {
			window.ShowAdd();
			return;
		}
		first = matches[0];
		setRadioValue("radType", getAttr(first, "MatType"));
		if (getAttr(first, "MatType") === "P") {
			selectByValue("cmbSCProcess", getAttr(first, "SCProcess"));
			selectByValue("cmbCurrency", getAttr(first, "Currency"));
			setValue("txtInstruct", getAttr(first, "Instruct"));
			setValue("txtLabourCharge", getAttr(first, "LabourCharge"));
		}
		matches.forEach(function (node) {
			if (getAttr(node, "MatRecdAsItem") || getAttr(node, "MatRecdAsCode")) {
				values.push(getAttr(node, "MatRecdAsItem") + ":" + getAttr(node, "MatRecdAsCode"));
				names.push(getAttr(node, "MatRecdAsDescr"));
				attrs.push(getAttr(node, "AttributeList"));
			}
		});
		if (names.length) {
			setSpanText("SpnMaterialToBeReceived", names.join(","));
			setValue("cmbMatRecdAs", values.join(","));
			setValue("hAttribute", attrs.join(","));
		}
		window.ShowAdd();
	};

	window.addEventListener("beforeunload", publishReturnValue);
})(window, document);
